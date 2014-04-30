VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEOAmend 
   Caption         =   "CWT TravelPro - EO Amend"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   9525
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdAmend 
      Caption         =   "&Amend"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtEONum 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin MSComctlLib.ListView lsvEO 
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "EO Number"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PC"
         Object.Width           =   1032
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Product"
         Object.Width           =   3792
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Last Amend Date"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Type"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "VendorCode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "EO Action"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lsvCompletedEO 
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "EO Number"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PC"
         Object.Width           =   1032
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Product"
         Object.Width           =   3792
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Completed Date"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Type"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "VendorCode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "EO Action"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblEONum 
      Caption         =   "EO Number"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Pending"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Completed"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   1815
   End
End
Attribute VB_Name = "frmEOAmend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datFormLoadStart As Date
Dim datFormLoadEnd As Date
Dim datTouchEnd As Date

Private Sub cmdAmend_Click()
 lsvEO_DblClick
End Sub

Private Sub cmdClose_Click()
'If fWantToQuit Then
    Unload Me
'    Call pRedisplayMenu
'End If
End Sub

Private Sub cmdSearch_Click()
   Dim strMsg As String
   
   If txtEONum = "" Then Exit Sub
   GetEOList
   If lsvEO.ListItems.Count = 0 And Me.lsvCompletedEO.ListItems.Count = 0 Then
      'MsgBox "Exchange order not found"
      strMsg = "Exchange order not found"
      modMsgBox.OKMsg = "OK"
      modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
   End If
End Sub

Private Sub Form_Load()
   Dim oldParent As Long
    preFormLoad
    datFormLoadStart = Now
    
    gintY = 0
    gintX = 0
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
    Me.Move 0, 0
    Me.Move frmSideBar.Width, 0
    GetEOList
    
     datFormLoadEnd = Now
    If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey

End Sub

Private Sub GetEOList()
   Dim strSql As String
   Dim rsEO As ADODB.Recordset
   
   'Set gobjPNR = New CWT_GalileoPNR.PNR
   'If gobjLog.LogOpen Then gobjPNR.OpenLog gobjLog
   Set gobjPNR = New CWT_GalileoPNR3.PNR
   gobjPNR.loadPNR

   strSql = "Select A.ExchangeID as ExID, A.ProductCode as PC, B.Description as Product," & _
   "A.CompletedDtTm as CompleteDt, A.LastAmendDtTm as AmendDt," & _
   "B.Type as Type, A.VendorCode as VendorCode, A.EOType as EOAction " & _
   "from tblExchangeOrder A, tblProductCodes B where "
   'strSQL = strSQL & "(CompletedDtTm is Null ) and "
   If txtEONum = "" Then
      If gobjPNR.RecLoc = "" Then Exit Sub
      strSql = strSql & "PNR = '" & gobjPNR.RecLoc & "' and "
   Else
      strSql = strSql & "ExchangeID = '" & txtEONum & "' and "
   End If
   strSql = strSql & "A.ProductCode = B.ProductCode "
   strSql = strSql & "Order by A.TransDate "

   Set rsEO = gdbConn.Execute(strSql)
   
   lsvEO.ListItems.Clear
   lsvCompletedEO.ListItems.Clear
   With rsEO
   Do Until .EOF
      If IsDate(!CompleteDt) Then
         Set item = lsvCompletedEO.ListItems.Add(, , !ExID)
      Else
         Set item = lsvEO.ListItems.Add(, , !ExID)
      End If
      item.SubItems(1) = !PC & ""
      item.SubItems(2) = !Product & ""
      If IsDate(!CompleteDt) Then
         item.SubItems(3) = !CompleteDt & ""
      Else
         item.SubItems(3) = !AmendDt & "" '!CreateDtTm & ""
      End If
      item.SubItems(4) = !Type & ""
      item.SubItems(5) = !VendorCode & ""
      item.SubItems(6) = !EOAction & ""
      .MoveNext
   Loop
   End With
End Sub


Private Sub lsvEO_DblClick()
   Dim objForm As Form
   Dim strFormName As String
   Dim strMsg As String
   
   datTouchEnd = Now
   If lsvEO.ListItems.Count = 0 Then Exit Sub
   
   frmOthSvcs.dbcProducts.BoundText = lsvEO.SelectedItem.SubItems(1)
   
   'frmOthSvcs.datProducts.DatabaseName = gstrTProDBSource
   'frmOthSvcs.datProducts.RecordSource = "SELECT * FROM tblProductCodes where ProductCode = '" & lsvEO.SelectedItem.SubItems(1) & "'"
   'frmOthSvcs.datProducts.Refresh
   With frmOthSvcs.datProducts.Recordset
      If Not .BOF Then .MoveFirst
      .Find "ProductCode = '" & lsvEO.SelectedItem.SubItems(1) & "'"
      If .EOF Then
         'MsgBox "Invalid product code: " & lsvEO.SelectedItem.SubItems(1)
         strMsg = "Invalid product code: " & lsvEO.SelectedItem.SubItems(1)
         modMsgBox.OKMsg = "OK"
         modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
         Exit Sub
      End If
   End With
   
   
   gbolEOAmend = True
   Select Case lsvEO.SelectedItem.SubItems(4)
     Case "HL"
         Set objForm = frmOSHotel
         strFormName = "frmOSHotel"
     Case "CT", "BT"
         Set objForm = frmOSAirTkt
         strFormName = "frmOSAirTkt"
     Case "CX"
         Set objForm = frmOSCarTxfr
         strFormName = "frmOSCarTxfr"
     Case "MS"
         Set objForm = frmOSMisc
         strFormName = "frmOSMisc"
     Case "VI"
         Set objForm = frmOSVisa
         strFormName = "frmOSVisa"
     Case "TR"
         Set objForm = frmOSOthTkt
         strFormName = "frmOSOthTkt"
     Case Else
         Exit Sub
     End Select
     'objForm.Left = Me.Left
     'objForm.Show
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModAux, _
       Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModAux, _
       Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModAux, _
       Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd
     
     
        Load objForm
        objForm.Show
        Do
            DoEvents
        Loop Until isLoaded(strFormName) = False
        
     Set objForm = Nothing
     
End Sub



