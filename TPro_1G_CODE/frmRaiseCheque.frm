VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRaiseCheque 
   Caption         =   "CWT TravelPro - Raise Cheque"
   ClientHeight    =   6390
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   11415
   Begin VB.CommandButton cmdSummary 
      Caption         =   "Summary Report"
      Height          =   495
      Left            =   6720
      TabIndex        =   20
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdVoid 
      Caption         =   "&Void"
      Height          =   495
      Left            =   4080
      TabIndex        =   19
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CheckBox chkCREO 
      Caption         =   "Print EO using crystal report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   17
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   495
      Left            =   5400
      TabIndex        =   16
      Top             =   5760
      Width           =   1215
   End
   Begin VB.ComboBox cmbListType 
      Height          =   315
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1200
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   5640
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCSV 
      Caption         =   "C&SV"
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtEONum 
      Height          =   375
      Left            =   1080
      MaxLength       =   20
      TabIndex        =   9
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtPNR 
      Height          =   375
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   2895
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   6240
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   8040
      TabIndex        =   4
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtChequeNum 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   5160
      Width           =   3015
   End
   Begin MSComctlLib.ListView lsvEO 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   27
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "PNR"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "EO Num"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Product"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Vendor"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Address1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Address2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "City"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Country"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "PxName"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Agent"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Tel"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Date"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Cost"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Tax1"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "TaxCode1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Tax2"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "TaxCode2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Description"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Remark"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "CN"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Fax"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Email"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "ContactPerson"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "AgentID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Cheque Num"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "Amend"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "ProductCode"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblTotalAmt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   15
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label lblTotal 
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   14
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label lblRaiseType 
      Caption         =   "Raise Type"
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblEONum 
      Caption         =   "EO Num"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblPNR 
      Caption         =   "PNR"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblChequeNum 
      Caption         =   "Cheque Num"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   5160
      Width           =   1095
   End
End
Attribute VB_Name = "frmRaiseCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sngEOTotal As Single
Dim mbolReportPath As Boolean
Dim mbolFormLoad As Boolean
Dim datTouchStart As Date
Dim datFormLoadStart As Date
Dim datTouchEnd As Date

Private Sub GetEOList(EOType As String)
   Dim strSql As String
   Dim rsEO As New ADODB.Recordset
   Dim objEO As New EO
   Dim item As ListItem
   Dim strVendorInfo() As String
   Dim i As Integer
   
   txtChequeNum = ""
   lsvEO.ListItems.Clear
   
   strSql = "Select A.NETTCOSTGST, A.PNR,A.VendorInfo,A.ExchangeID,A.Name,A.Description as EODesc, " & _
            "AgentPhone,A.CompletedDtTm,A.LastAmendDtTm,A.Cost,A.Tax1,A.Tax2,A.ChequeNum, " & _
            "A.CreatedBy,A.TaxCode1,A.TaxCode2,A.Remarks,A.CN,A.ContactPerson, " & _
            "B.Description as Product,C.VendorName,C.Address1,C.Address2,C.City, " & _
            "C.Country,C.Email,C.FaxNumber,C.Misc,D.AgentName,A.ProductCode,C.RaiseType " & _
            "from tblExchangeOrder A, tblProductCodes B, tblVendors C, tblAgents D where "

   If EOType = "RaiseCheque" Then
      strSql = strSql & "Finance = 'True' and (CompletedDtTm is Null ) and "
      strSql = strSql & "A.ProductCode = B.ProductCode and A.VendorCode = C.VendorNumber and "
      strSql = strSql & "Void =0 and "
      'strSQL = strSQL & "A.CreatedBy = D.Sine "
      strSql = strSql & "A.LastAmendBy = D.Sine and A.LastAmendByPCC = D.PCC "
      strSql = strSql & "and C.RaiseType='" & cmbListType & "' "
      strSql = strSql & "Order by VendorName, A.TransDate "
   
   ElseIf EOType = "ReprintEO" Then
      'strSQL = strSQL & "Finance = True and (CompletedDtTm is Null ) and "
      strSql = strSql & "A.PNR = '" & gobjPNR.RecLoc & "' and "
      strSql = strSql & "A.ProductCode = B.ProductCode and A.VendorCode = C.VendorNumber and "
      strSql = strSql & "Void =0 and "
      'strSQL = strSQL & "A.CreatedBy = D.Sine "
      strSql = strSql & "A.LastAmendBy = D.Sine and A.LastAmendByPCC = D.PCC "
      strSql = strSql & "Order by A.TransDate "
   Else
      strSql = strSql & "Finance = 'True' and (CompletedDtTm is not Null ) and "
      If Trim(txtPNR) <> "" Then
         strSql = strSql & "PNR = '" & Trim(txtPNR) & "' and "
      End If
      If Trim(txtEONum) <> "" Then
         strSql = strSql & "ExchangeID = '" & Trim(txtEONum) & "' and "
      End If
      If Trim(txtPNR) = "" And Trim(txtEONum) = "" Then
         strSql = strSql & "CompletedDtTm >= '" & DateAdd("d", -15, Date) & "' and "
      End If
      strSql = strSql & "A.ProductCode = B.ProductCode and A.VendorCode = C.VendorNumber and "
      'strSql = strSql & "Void = False and "
      'strSQL = strSQL & "A.CreatedBy = D.Sine "
      strSql = strSql & "A.LastAmendBy = D.Sine and A.LastAmendByPCC = D.PCC "
      strSql = strSql & "Order by A.TransDate "
   End If
   
   If EOType = "Amend" Then
      lsvEO.ColumnHeaders(12).Text = "Completed Time"
   Else
      lsvEO.ColumnHeaders(12).Text = "Last Amend Time"
   End If
   
  'strSQL = "Select A.PNR,A.VendorInfo,A.ExchangeID, AgentPhone,A.CompletedDtTm,A.LastAmendDtTm," & _
  '          "A.Cost,A.Tax1,A.Tax2,A.ChequeNum," & _
  '          "A.CreatedBy,A.TaxCode1,A.TaxCode2,A.Remarks,A.CN,A.ContactPerson, " & _
  '          "C.VendorName,C.Address1,C.Address2,C.City, " & _
  '          "C.Country , C.Email, C.FaxNumber, C.Misc, D.AgentName " & _
  '          "from tblExchangeOrder A, tblProductCodes B, tblVendors C, tblAgents D " & _
  '          "where a.exchangeid= '8850503000300' and A.PNR = 'LX68CW' and A.ProductCode = B.ProductCode and A.VendorCode = C.VendorNumber" & _
  '          " and A.LastAmendBy = D.Sine" & _
  '          " Order by A.TransDate"

   
   'Set rsEO = gdbConn.Execute(strSQL)
   rsEO.Open strSql, gdbConn, adOpenKeyset, adLockReadOnly
   With rsEO
   Do Until .EOF
      'If !VendorCode = "999999" Then
      If !Misc = "True" And !VendorInfo <> "" Then
         strVendorInfo = Split(!VendorInfo, vbCrLf)
      Else
         ReDim strVendorInfo(6)
         strVendorInfo(0) = !VendorName & ""
         strVendorInfo(1) = !Address1 & ""
         strVendorInfo(2) = !Address2 & ""
         strVendorInfo(3) = !City & ""
         strVendorInfo(4) = !Country & ""
         strVendorInfo(5) = !Email & ""
         strVendorInfo(6) = !FaxNumber & ""
      End If
   
      Set item = lsvEO.ListItems.Add(, , !PNR)
      item.SubItems(1) = !ExchangeID
      item.SubItems(2) = !Product & ""
      item.SubItems(3) = strVendorInfo(0) '!VendorName & ""
      item.SubItems(4) = strVendorInfo(1) '!Address1 & ""
      item.SubItems(5) = strVendorInfo(2) '!Address2 & ""
      item.SubItems(6) = strVendorInfo(3) '!City & ""
      item.SubItems(7) = strVendorInfo(4) '!Country & ""
      item.SubItems(8) = !Name & ""
      item.SubItems(9) = !AgentName & ""
      item.SubItems(10) = !AgentPhone & ""
      If EOType = "Amend" Then
         item.SubItems(11) = !CompletedDtTm '!CreateDtTm
      Else
         item.SubItems(11) = !LastAmendDtTm '!CreateDtTm
      End If
      item.SubItems(12) = Format(!Cost, "0.00")
      'Item.SubItems(13) = !Tax1
      'Item.SubItems(14) = !TaxCode1 & ""
       If !TAXCODE1 = "G*" Then
        If !NETTCOSTGST > 0 Then
          item.SubItems(13) = !NETTCOSTGST
        Else
          item.SubItems(13) = !Tax1
        End If
          item.SubItems(14) = !TAXCODE1 & ""
      Else
        item.SubItems(13) = !Tax1
        item.SubItems(14) = !TAXCODE1 & ""
      End If
      item.SubItems(15) = !Tax2 & ""
      item.SubItems(16) = !TaxCode2 & ""
      item.SubItems(17) = !EODesc & ""
      item.SubItems(18) = !Remarks & ""
      item.SubItems(19) = !CN & ""
      item.SubItems(20) = strVendorInfo(6) '!FaxNumber & ""
      item.SubItems(21) = strVendorInfo(5) '!Email & ""
      item.SubItems(22) = !ContactPerson & ""
      item.SubItems(23) = !CreatedBy & ""
      item.SubItems(24) = !ChequeNum & ""
      If !CompletedDtTm <> !LastAmendDtTm Then
         item.SubItems(25) = "True"
      Else
         item.SubItems(25) = "False"
      End If
      item.SubItems(26) = !ProductCode & ""
      rsEO.MoveNext
   Loop
   End With
   rsEO.Close
   Set rsEO = Nothing
   If lsvEO.ListItems.Count <> 0 Then
      'lsvEO.ListItems(1).Selected = True
      txtChequeNum = lsvEO.SelectedItem.SubItems(24)
   End If
End Sub


Private Sub chkCREO_Click()
   Dim strMsg As String
   
   If chkCREO.value = 1 Then
      If mbolReportPath Then
      
      Else
          chkCREO.value = 0
          'MsgBox "Report Path not found."
          strMsg = "Report Path not found."
          modMsgBox.OKMsg = "OK"
          modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
      End If
   End If
End Sub


Private Sub cmbListType_Click()
    GetEOList ("RaiseCheque")
         sngEOTotal = 0
         lblTotalAmt.Caption = Format(sngEOTotal, "0.00")

End Sub

Private Sub cmbType_Click()
   'If cmbType.Visible Then
   If gbolReprintEO = False Then
      If cmbType.Text = "Raise Cheque" Then
         lsvEO.ColumnHeaders.item(23).Width = 1440
         lsvEO.ColumnHeaders.item(25).Width = 0
         lblRaiseType.Visible = True
         cmbListType.Visible = True
         PopulateRaiseType
         'If mbolFormLoad = False Then
         '   GetEOList "RaiseCheque"
         'End If
         txtPNR.Enabled = False
         txtEONum.Enabled = False
         lblTotalAmt.Visible = True
         lblTotal.Visible = True
         sngEOTotal = 0
         lblTotalAmt.Caption = Format(sngEOTotal, "0.00")
         cmdReset.Visible = False
      Else
         lsvEO.ColumnHeaders.item(25).Width = 1440
         lsvEO.ColumnHeaders.item(23).Width = 0
         txtPNR = ""
         txtEONum = ""
         lblRaiseType.Visible = False
         cmbListType.Visible = False
         
         lsvEO.ListItems.Clear
         txtPNR.Enabled = True
         txtEONum.Enabled = True
         
         lblTotalAmt.Visible = False
         lblTotal.Visible = False
         GetEOList "Amend"
         cmdReset.Visible = True
      End If
   Else
      lsvEO.ColumnHeaders.item(23).Width = 1440
      lsvEO.ColumnHeaders.item(25).Width = 0
      GetEOList "ReprintEO"
   End If
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdCSV_Click()
   Dim strFile As String
   Dim i As Integer
   Dim strText As String
   Dim strMsg As String
   datTouchEnd = Now
   Dialog1.Filter = "Text (*.txt)|*.txt|CSV (*.csv)|*.csv"
   Dialog1.ShowSave
   strFile = Dialog1.FileName
   If strFile = "" Then Exit Sub
   
   Open strFile For Output As #1
  For i = 1 To lsvEO.ListItems.Count
     
      With lsvEO.ListItems(i)
         strText = """" & .SubItems(1) & """,""" & .SubItems(11) & """,""" & .SubItems(8) & """,""" & .Text & """,""" & .SubItems(13) & """,""" & .SubItems(23) & """"
         Print #1, strText
      End With
  Next
  Close #1
  Dialog1.FileName = ""
  'MsgBox "Done"
  strMsg = "Done"
  modMsgBox.OKMsg = "OK"
  modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop"
  EndLog
  datTouchStart = Now

End Sub

Private Sub cmdPrint_Click()
   Dim intI As Integer
   Dim strEONum As String
   datTouchEnd = Now
   If lsvEO.ListItems.Count = 0 Then Exit Sub
   
'   If cmdUpdate.Caption = "Reprint" Then
'       For intI = 1 To lsvEO.ListItems.Count
'        If lsvEO.ListItems(intI).Selected Then
'            If chkCREO.Value = 0 Then
'               PrintEO (intI)
'            Else
'               'PrintCREO lsvEO.SelectedItem.SubItems(1), intI
'               strEONum = strEONum & IIf(strEONum = "", "'" & lsvEO.ListItems(intI).SubItems(1), "," & "'" & lsvEO.ListItems(intI).SubItems(1)) & "'"
'            End If
'        End If
'       Next
'
'       If chkCREO.Value = 1 Then
'                gEOID = strEONum
'                Load frmEOCrystalReport
'                frmEOCrystalReport.Show 1, Me
'       End If
'print
'      Exit Sub
'   End If
      
   For intI = 1 To lsvEO.ListItems.Count
    If lsvEO.ListItems(intI).Selected Then
         'strSql = "UPDATE tblExchangeOrder set CompletedBy='', CompletedDtTm=getDate()," & _
         '         "ChequeNum='" & txtChequeNum & "' where ExchangeID = '" & lsvEO.ListItems(intI).SubItems(1) & "'"
         'gdbConn.Execute (strSql)
         If chkCREO.value = 0 Then
            PrintEO (intI)
         Else
            strEONum = strEONum & IIf(strEONum = "", "'" & lsvEO.ListItems(intI).SubItems(1), "," & "'" & lsvEO.ListItems(intI).SubItems(1)) & "'"
            'strEONum = strEONum & "'" & IIf(strEONum = "", lsvEO.ListItems(intI).SubItems(1), "," & lsvEO.ListItems(intI).SubItems(1)) & "'"
         End If
    End If
   Next
   If chkCREO.value = 1 Then
    gEOID = strEONum
    'gEOReportName = "EO.rpt"
    '  Load frmEOCrystalReport
    '  frmEOCrystalReport.Show 1, Me
      
   gbolPreviewEO = False
   gbolRaiseEOReport = True
   gbolIndEO = True
   gEOReportName = "EO" & "-" & gstrAgcyCountryCode & ".rpt"
   
   EndLog
   
   
   Load frmEOCrystalReport
   'frmEOCrystalReport.mnuAccept.Caption = "Print"
   frmEOCrystalReport.mnuAccept.Visible = False
   frmEOCrystalReport.mnuCancel.Caption = "Close"
   frmEOCrystalReport.Caption = "EO"
   frmEOCrystalReport.Show
   Do
     DoEvents
   Loop Until isLoaded("frmEOCrystalReport") = False
   gbolRaiseEOReport = False
   gbolIndEO = False
   End If
   
   datTouchStart = Now
   
End Sub

Private Sub cmdRefresh_Click()
   cmbType_Click
End Sub

Private Sub cmdReset_Click()
Dim intI As Integer
Dim strSql As String
datTouchEnd = Now
   For intI = 1 To lsvEO.ListItems.Count
    If lsvEO.ListItems(intI).Selected Then
         strSql = "UPDATE tblExchangeOrder set CompletedBy=null, CompletedDtTm=null" & _
                  " ,Void = 0" & _
                  " where ExchangeID = '" & lsvEO.ListItems(intI).SubItems(1) & "'"
         gdbConn.Execute (strSql)
         cmbType_Click
         Exit For
    End If
   Next
EndLog
datTouchStart = Now
End Sub

Private Sub cmdSummary_Click()
   Dim intI As Integer
   Dim strEONum As String
   Dim strMsg As String
   datTouchEnd = Now
   If lsvEO.ListItems.Count = 0 Then Exit Sub
   For intI = 1 To lsvEO.ListItems.Count
    If lsvEO.ListItems(intI).Selected Then
       strEONum = strEONum & IIf(strEONum = "", "'" & lsvEO.ListItems(intI).SubItems(1), "," & "'" & lsvEO.ListItems(intI).SubItems(1)) & "'"
    End If
   Next
   If strEONum = "" Then
      'MsgBox "Please select exchange order"
      strMsg = "Please select exchange order"
      modMsgBox.OKMsg = "OK"
      modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
      Exit Sub
   End If
   gEOID = strEONum
   gbolPreviewEO = False
   gbolRaiseEOReport = True
   gEOReportName = "EOSummary" & "-" & gstrAgcyCountryCode & ".rpt"
   
   EndLog
   
   Load frmEOCrystalReport
   'frmEOCrystalReport.mnuAccept.Caption = "Print"
   frmEOCrystalReport.mnuAccept.Visible = False
   frmEOCrystalReport.Caption = "Summary EO"
   frmEOCrystalReport.Show
   Do
     DoEvents
   Loop Until isLoaded("frmEOCrystalReport") = False
   gbolRaiseEOReport = False
   datTouchStart = Now
End Sub

Private Sub cmdUpdate_Click()
   Dim strSql As String
   Dim intI As Integer
   
   'Dim rsEO As New ADODB.Recordset
   datTouchEnd = Now
   If lsvEO.ListItems.Count = 0 Then Exit Sub
   
   'If cmdUpdate.Caption = "Reprint" Then
   '    For intI = 1 To lsvEO.ListItems.Count
   '     If lsvEO.ListItems(intI).Selected Then
   '         If chkCREO.Value = 0 Then
   '            PrintEO (intI)
   '         Else
   '            'PrintCREO lsvEO.SelectedItem.SubItems(1), intI
   '            strEONum = strEONum & IIf(strEONum = "", "'" & lsvEO.ListItems(intI).SubItems(1), "," & "'" & lsvEO.ListItems(intI).SubItems(1)) & "'"
   '         End If
   '     End If
   '    Next
   '
   '    If chkCREO.Value = 1 Then
   '             gEOID = strEONum
   '             Load frmEOCrystalReport
   '             frmEOCrystalReport.Show 1, Me
   '    End If
   '
   '   Exit Sub
   'End If
      
   
   ''strSQL = "Select * from tblExchangeOrder where ExchangeID = '" & _
   ''         lsvEO.SelectedItem.SubItems(1) & "'"
   For intI = 1 To lsvEO.ListItems.Count
    If lsvEO.ListItems(intI).Selected Then
         strSql = "UPDATE tblExchangeOrder set CompletedBy='', CompletedDtTm=getDate()," & _
                  "ChequeNum='" & txtChequeNum & "' where ExchangeID = '" & lsvEO.ListItems(intI).SubItems(1) & "'"
         gdbConn.Execute (strSql)
   '      If chkCREO.Value = 0 Then
   '         PrintEO (intI)
   '      Else
   '         strEONum = strEONum & "'" & IIf(strEONum = "", lsvEO.ListItems(intI).SubItems(1), "," & lsvEO.ListItems(intI).SubItems(1)) & "'"
   '      End If
    End If
   Next
   '
   'If chkCREO.Value = 1 Then
   ' gEOID = strEONum
   '   Load frmEOCrystalReport
   '   frmEOCrystalReport.Show 1, Me
   'End If
   
   'rsEO.Edit
   'rsEO!CompletedBy = ""
   'rsEO!CompletedDtTm = Now
   'rsEO!ChequeNum = txtChequeNum
   'rsEO.Update
   
   'rsEO.Close
   'Set rsEO = Nothing
   
   
   If cmbType.Text = "Raise Cheque" Then
      GetEOList "RaiseCheque"
   Else
      lsvEO.ListItems.Clear
      txtChequeNum = ""
   End If
   EndLog
   datTouchStart = Now
End Sub

Private Sub cmdVoid_Click()
Dim intI As Integer
Dim strSql As String
   datTouchEnd = Now
   For intI = 1 To lsvEO.ListItems.Count
    If lsvEO.ListItems(intI).Selected Then
         strSql = "UPDATE tblExchangeOrder set Void = 1" & _
                  " where ExchangeID = '" & lsvEO.ListItems(intI).SubItems(1) & "'"
         gdbConn.Execute (strSql)
         cmbType_Click
         Exit For
    End If
   Next
   EndLog
   datTouchStart = Now
   
End Sub

Private Sub Form_Load()
   Dim oldParent As Long
   datFormLoadStart = Now
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
   Me.Move 0, 0
   Me.Move frmSideBar.Width, 0
mbolFormLoad = True
If gstrReportPath <> "" Then mbolReportPath = True
    If UCase(gstrAgcyCountryCode) = "SG" Then
        'chkCREO.Visible = True
        cmdSummary.Visible = True
    Else
        'chkCREO.Visible = False
        cmdSummary.Visible = False
    End If
    chkCREO.Visible = False
    chkCREO.value = vbChecked
   If gbolReprintEO Then
     Me.Caption = "EO Reprint"
     cmbType.Visible = False
     lblPNR.Visible = False
     txtPNR.Visible = False
     lblEONum.Visible = False
     txtEONum.Visible = False
     lblChequeNum.Visible = False
     txtChequeNum.Visible = False
     lblRaiseType.Visible = False
     cmbListType.Visible = False
     'cmdUpdate.Caption = "Reprint"
     cmdUpdate.Enabled = False
     cmdVoid.Enabled = False
     Set gobjPNR = New CWT_GalileoPNR3.PNR
     gobjPNR.loadPNR
     cmbType_Click
     lblTotal.Visible = False
     lblTotalAmt.Visible = False
     cmdReset.Visible = False
     cmdReset.Visible = False
   Else
     cmbType.Clear
     cmbType.AddItem "Raise Cheque"
     cmbType.AddItem "Amend Cheque Num."
     cmbType.listindex = 0
     Me.Caption = "Raise Cheque"
     
   End If
    mbolFormLoad = False
    datTouchStart = Now
    If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey
    StartLog
End Sub

Private Sub PrintEO(Optional listindex As Integer)
   Dim aryDesc() As String
   Dim aryRemark() As String
   Dim strTemp As String
   Dim intPrint As Integer
   Dim intI As Integer
   Dim intC As Integer
   Dim i As Integer
   
   Dialog.ShowPrinter
   
   'If gbolReprintEO = False Then
   ' intPrint = 2
   'Else
   ' intPrint = 1
   'End If
   
   'For intI = 1 To intPrint
   
   With lsvEO.ListItems(listindex)
   
   Printer.FontSize = 12
   Printer.FontBold = False
   Printer.FontName = "Courier new"
   Printer.Print ""
   Printer.Print ""
   For intC = 0 To 7
       Printer.Print ""
   Next
   If .SubItems(25) = True Then
      Printer.FontBold = True
      Printer.Print Space(4) & UCase("Amendment")
      Printer.FontBold = False
   End If
   If .SubItems(22) <> "" Then
      Printer.Print Space(4) & "Attn: " & .SubItems(22) 'Attn
   End If
   Printer.Print Space(4) & .SubItems(3) 'Vendor Name
   Printer.Print Space(4) & .SubItems(4) 'Addr1
   Printer.Print Space(4) & .SubItems(5) 'Addr2
   Printer.Print Space(4) & .SubItems(6) 'City
   Printer.Print Space(4) & .SubItems(7) 'Country
   Printer.Print ""
   Printer.FontBold = True
   Printer.Print Space(4) & "EXCHANGE ORDER"
   Printer.Print Space(4) & "--------------"
   Printer.FontBold = False
   Printer.Print Space(4) & "EO Number      : " & .SubItems(1)
   Printer.Print Space(4) & "Passenger Name : " & .SubItems(8)
   Printer.Print Space(4) & "Agent Name     : " & .SubItems(9)
   Printer.Print Space(4) & "Agent ID       : " & .SubItems(23)
   Printer.Print Space(4) & "Record Locator : " & .Text
   Printer.Print Space(4) & "TEL            : " & .SubItems(10)
   'Printer.Print Space(4) & "To             : " & gobjEO.VendorName
   'Printer.Print Space(4) & "Fax            : " & gobjEO.FaxNo
   'Printer.Print Space(4) & "Email          : " & gobjEO.Email
   Printer.Print Space(4) & "Date           : " & Format(.SubItems(11), "Medium Date")
   Printer.Print ""
   Printer.FontBold = True
   Printer.Print Space(4) & "Service Info  "
   Printer.FontBold = False
   Printer.Print Space(4) & "Nett Cost      : " & .SubItems(12)
   If .SubItems(26) = "02" Then
        If UCase(.SubItems(14)) = "G*" Then                             'TaxCode1
              Printer.Print Space(4) & "GST            : " & Format(.SubItems(13), gstrAgcyCurrFormat) 'Tax1
           ElseIf UCase(.SubItems(16)) = "G*" Then                         'TaxCode2
              Printer.Print Space(4) & "GST            : " & Format(.SubItems(15), gstrAgcyCurrFormat) 'Tax2
        End If
        If UCase(.SubItems(14)) <> "G*" Then                             'TaxCode1
            Printer.Print Space(4) & "Tax            : " & Format(.SubItems(13), gstrAgcyCurrFormat) & .SubItems(14) 'Tax1 & Code
        End If
        If UCase(.SubItems(16)) <> "G*" Then                             'TaxCode2
           Printer.Print Space(4) & "Tax            : " & Format(.SubItems(15), gstrAgcyCurrFormat) & .SubItems(16)
        End If
   Printer.Print Space(4) & "Total          : " & Format(fConvertZero(.SubItems(12)) + fConvertZero(.SubItems(13)) + fConvertZero(.SubItems(15)), gstrAgcyCurrFormat)
   End If
   
        
  'remove on 21/4/05: requested by helena
  
  ' ElseIf gobjEO.TaxCount = 1 Then
  '    If UCase(gobjEO.Tax(1).Code) = "GST" Then
  '       Printer.Print Space(4) & "GST            : " & gobjEO.Tax(1).Amount
  '    Else
  '       Printer.Print Space(4) & "Tax            : " & gobjEO.Tax(1).Amount & gobjEO.Tax(1).Code
  '    End If
  
   
   'Printer.Print Space(4) & "Commission     : " & gobjEO.CommissionAmt
   'If .SubItems(14) <> "" And .SubItems(16) <> "" Then
   '   If UCase(.SubItems(14)) = "G*" Then                             'TaxCode1
   '      Printer.Print Space(4) & "GST            : " & .SubItems(13) 'Tax1
   '   ElseIf UCase(.SubItems(16)) = "G*" Then                         'TaxCode2
   '      Printer.Print Space(4) & "GST            : " & .SubItems(15) 'Tax2
   '   End If
   '   If UCase(.SubItems(14)) = "G*" Then                             'TaxCode1
   '      Printer.Print Space(4) & "Tax            : " & .SubItems(13) & .SubItems(14) 'Tax1 & Code
   '   End If
   '   If UCase(.SubItems(16)) = "G*" Then                             'TaxCode2
   '      Printer.Print Space(4) & "Tax            : " & gobjEO.Tax(2).Amount & gobjEO.Tax(2).Code
   '   End If
   'ElseIf .SubItems(14) <> "" Then
   '   If UCase(.SubItems(14)) = "G*" Then                             'TaxCode1
   '      Printer.Print Space(4) & "GST            : " & .SubItems(13)
   '  Else
   '      Printer.Print Space(4) & "Tax            : " & .SubItems(13) & .SubItems(14)
   '   End If
   'End If
   
 '02062005
   'aryDesc = Split(.SubItems(17), ";")
   'aryRemark = Split(.SubItems(18), ";")
   If InStr(1, .SubItems(17), vbCrLf) <> 0 Then
      aryDesc = Split(.SubItems(17), vbCrLf)
   Else
   aryDesc = Split(.SubItems(17), ";")
   End If
   If InStr(1, .SubItems(18), vbCrLf) <> 0 Then
      aryRemark = Split(.SubItems(18), vbCrLf)
   Else
   aryRemark = Split(.SubItems(18), ";")
   End If

   For i = 0 To UBound(aryDesc)
      If i = 0 Then
         Printer.Print Space(4) & "Description    : " & aryDesc(i)
      Else
         Printer.Print Space(4) & "               " & aryDesc(i)
      End If
   Next
   If UBound(aryRemark) <> -1 Then
      For i = 0 To UBound(aryRemark)
         If i = 0 Then
            Printer.Print Space(4) & "Remark         : "
            Printer.Print Space(4) & aryRemark(i)
         Else
            Printer.Print Space(4) & aryRemark(i)
         End If
      Next
   End If
   Printer.Print ""
   Printer.Print Space(4) & "Please prepare document for our collection today."
   Printer.Print Space(4) & "Thank you"
   Printer.EndDoc
   End With
   'MsgBox "Print Successful!", vbOKOnly, "Print Successful"
   
   'Next
   
End Sub

Private Sub PopulateRaiseType()
   Dim rsRT As ADODB.Recordset
   Dim strSql As String
   
   
   
   cmbListType.Clear
   
   strSql = "SELECT DISTINCT(RaiseType) from tblVendors"
   Set rsRT = gdbConn.Execute(strSql)
   While Not rsRT.EOF
   If Not IsNull(rsRT!RaiseType) Then
        cmbListType.AddItem rsRT!RaiseType
    End If
        rsRT.MoveNext

   Wend
   
   If cmbListType.ListCount > 1 Then cmbListType.listindex = 0
   
End Sub

Private Sub lsvEO_Click()
Dim intI As Integer
sngEOTotal = 0
lblTotalAmt.Caption = Format(sngEOTotal, "0.00")

For intI = 1 To lsvEO.ListItems.Count
If lsvEO.ListItems(intI).Selected = True Then
   With lsvEO.ListItems(intI)
   sngEOTotal = sngEOTotal + fConvertZero(.SubItems(12)) + fConvertZero(.SubItems(13)) + fConvertZero(.SubItems(15))
    End With

End If
  Next
   lblTotalAmt.Caption = Format(sngEOTotal, "0.00")

End Sub

Private Sub lsvEO_DblClick()
 If UCase(gstrAgcyCountryCode) = "SG" Then
    'chkCREO.value = 1
    cmdPrint_Click
    'chkCREO.value = 0
 End If
End Sub

Private Sub lsvEO_ItemClick(ByVal item As MSComctlLib.ListItem)
Dim intI As Integer
Dim intSelCount As Integer


If Not gbolReprintEO Then
    'If cmbType.Text <> "Raise Cheque" Then
       If lsvEO.ListItems.Count > 0 Then
        intSelCount = 0
           For intI = 1 To lsvEO.ListItems.Count
                If intSelCount > 2 Then Exit For
                If lsvEO.ListItems(intI).Selected = True Then
                   intSelCount = intSelCount + 1
                End If
                
           Next intI
        If intSelCount < 2 Then
           If lsvEO.SelectedItem.SubItems(24) <> "" Then
             txtChequeNum = lsvEO.SelectedItem.SubItems(24)
           Else
             txtChequeNum = ""
           End If
        End If
       End If
    'End If
End If
End Sub

Private Sub txtEONum_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And Trim(txtEONum) <> "" Then GetEOList "Amend"
End Sub

Private Sub txtPNR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And Trim(txtPNR) <> "" Then GetEOList "Amend"
End Sub

Private Sub EndLog()
             
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, IIf(gbolReprintEO = True, gconSModReprintEO, gconSModApproveChq), _
       Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datTouchStart
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, IIf(gbolReprintEO = True, gconSModReprintEO, gconSModApproveChq), _
       Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd
End Sub
Private Sub StartLog()

       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, IIf(gbolReprintEO = True, gconSModReprintEO, gconSModApproveChq), _
       Me.Name, gconFormLoad, gstrProcessGrpID, datTouchStart, datFormLoadStart

End Sub
