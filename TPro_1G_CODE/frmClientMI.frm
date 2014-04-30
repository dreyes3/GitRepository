VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmClientMI 
   Caption         =   "CWT TravelPro - Client MI Entry"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   8295
   Begin VB.PictureBox cmbContainer 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5280
      ScaleHeight     =   375
      ScaleWidth      =   1005
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   1000
      Begin MSForms.ComboBox cmbClientMI 
         Height          =   375
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   855
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "1508;661"
         ListWidth       =   7055
         ColumnCount     =   2
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
         Object.Width           =   "1762;7055"
      End
   End
   Begin VB.TextBox txtDataModified 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3720
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbMICat 
      Height          =   315
      ItemData        =   "frmClientMI.frx":0000
      Left            =   1440
      List            =   "frmClientMI.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   4800
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvwClientMI 
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5530
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "FF Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "FF Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "FF Value"
         Object.Width           =   3625
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Data Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Data Length"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "DI Num"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Category:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Client MI Entry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "frmClientMI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public intLocation As Integer
Public intFileFareNum As Integer
Public intProdCode As String
Public intPaxID As Integer
Public MSXfreefields As String
Public DIfreefields As String
'Preethi - V1.2.2 20110301 -  IR8  - To solve the issue where the consultants can bypass the MI input
'Public bolCheck As Boolean
'230108
Public strPdtType As String
'
'The following objects are for multi-line list view entry
'
Private Type LSTVIEWITEM
  mask As Long
  lngItem As Long
  lngSubItem As Long
  state As Long
  stateMask As Long
  pszText As String
  cchTextMax As Long
  lngImage As Long
  lngParam As Long
  lngIndent As Long
End Type

Private dataSetup As Boolean
Private dataModified As Boolean
Private itmClicked As ListItem
Private dwLastSubitemEdited As Long

Private Const LVM_FIRST = &H1000
Private Const LVM_GETCOLUMNWIDTH = (LVM_FIRST + 29)
Private Const LVM_GETTOPINDEX = (LVM_FIRST + 39)
Private Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
Private Const LVM_SUBITEMHITTEST = (LVM_FIRST + 57)
Private Const LVHT_ONITEMICON = &H2
Private Const LVHT_ONITEMLABEL = &H4
Private Const LVHT_ONITEMSTATEICON = &H8
Private Const LVHT_ONITEM = (LVHT_ONITEMICON Or _
                           LVHT_ONITEMLABEL Or _
                           LVHT_ONITEMSTATEICON)
Private Const LVIR_LABEL = 2

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type LVHITTESTINFO
  pt As POINTAPI
  flags As Long
  lngItem As Long
  lngSubItem  As Long
End Type

Private Declare Function ScreenToClient Lib "user32" _
  (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Declare Function SendMessage Lib "user32" _
  Alias "SendMessageA" _
(ByVal hwnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  lngParam As Any) As Long
'
' end declaration for multi-line list view entry
'
'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
Public colFFValue As Collection

'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI

Private Sub cmbClientMI_Change()
If Not dataSetup Then
   dataModified = True
End If
End Sub

Private Sub cmbClientMI_KeyPress(KeyCode As MSForms.ReturnInteger)
KeyCode = Asc(UCase(Chr(CInt(KeyCode))))
End Sub

Private Sub cmbClientMI_LostFocus()
If dataModified And dwLastSubitemEdited > 0 Then
   itmClicked.SubItems(dwLastSubitemEdited) = cmbClientMI.Text
   dataModified = False
End If

End Sub

Private Sub cmbClientMI_Validate(Cancel As Boolean)
Dim strMsg As String
If IsNumeric(lvwClientMI.SelectedItem.Index) Then
    If Not validMI(lvwClientMI.SelectedItem.Index) Then
        'MsgBox "Invalid data/Incompatible length detected.", vbCritical
        strMsg = "Invalid data/Incompatible length detected."
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
        cmbClientMI.SetFocus
    End If
End If
End Sub

Private Sub cmbMICat_click()
    Dim strSql As String
    Dim strSQL2 As String
    
    If UCase(cmbMICat.Text) = "ALL" Then
        strSql = "SELECT tblClientMI.*, tblMICategory.format FROM tblClientMI, tblMICategory where cn = '" & gobjPNR.CN & "' and tblClientMI.location = tblMICategory.code and isNumeric(FF)= 1 order by cast(FF as integer)"
        strSQL2 = "SELECT tblClientMI.*, tblMICategory.format FROM tblClientMI, tblMICategory where cn = '" & gobjPNR.CN & "' and tblClientMI.location = tblMICategory.code and isNumeric(FF)= 0 order by FF"
    Else
        strSql = "SELECT tblClientMI.*, tblMICategory.format FROM tblClientMI, tblMICategory where cn = '" & gobjPNR.CN & "'" & _
                 " and location = " & cmbMICat.listindex & "  and tblClientMI.location = tblMICategory.code and isNumeric(FF)= 1 order by cast(FF as integer)"
        strSQL2 = "SELECT tblClientMI.*, tblMICategory.format FROM tblClientMI, tblMICategory where cn = '" & gobjPNR.CN & "'" & _
                 " and location = " & cmbMICat.listindex & "  and tblClientMI.location = tblMICategory.code and isNumeric(FF)= 0 order by FF"
    End If
    pRefreshData strSql, True
    pRefreshData strSQL2, False
End Sub

Private Sub cmdCancel_Click()
If pGetMIFormat(cmbMICat.listindex) = "D" Then
    Unload Me
    'pRedisplayMenu
Else
    Unload Me
End If
End Sub

Private Sub cmdDone_Click()
    Dim strMsg As String
    
    'If incompleteMI Then
    '    If Not MsgBox("Are you sure to exit without entering complete MI data?", vbYesNo) = vbYes Then
    '        bolCheck = True
    '        Exit Sub
    '    End If
    'End If
    strMsg = incompleteMI
    If strMsg <> "" Then
       'MsgBox strMsg, vbCritical
       modMsgBox.OKMsg = "OK"
       modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
       Exit Sub
    End If
    
    Select Case pGetMIFormat(cmbMICat.listindex)
        Case "D"
            DIfreefields = getDIFreeFields
            If cmbMICat.listindex = 3 Or cmbMICat.listindex = 7 Then  'Pre-Trip MI
                updateDIField DIfreefields
                gobjHost.terminalEntry DIfreefields
                gobjHost.terminalEntry "R.TPRO CLIENTMI"
                gobjHost.terminalEntry "ER"
                gobjHost.terminalEntry "ER"
                gobjHost.terminalEntry "ER"
            End If
            Unload Me
        Case "F"
             '230108
            If strPdtType = "BT" Or strPdtType = "CT" Or _
            intProdCode = "35" Or _
            intProdCode = "41" Or _
            intProdCode = "50" Or _
            intProdCode = "70" Then
                MSXfreefields = getMSXFreeFields
                Me.Hide
            Else
                DIfreefields = getDIFreeFields
                Me.Hide
            End If
        Case "M"
            MSXfreefields = getMSXFreeFields
            Me.Hide
        Case Else
            Unload Me
    End Select
End Sub


Private Sub Form_Load()
        
' This code is for multi-line list view entry
'
  Dim itmx As ListItem
  Dim cnt As Long
  Dim oldParent As Long
    
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
  Me.Move 0, 0
  Me.Move frmSideBar.Width, 0
  
  With lvwClientMI
     '.ColumnHeaders.Add , , "col 1"
     '.ColumnHeaders.Add , , "col 2"
     '.ColumnHeaders.Add , , "col 3"
     '.ColumnHeaders.Add , , "col 4"

     'For cnt = 1 To 20
     '   Set itmx = .ListItems.Add(, , "Item " & CStr(cnt))
     '   itmx.SubItems(1) = "subitem 1," & CStr(cnt)
     '   itmx.SubItems(2) = "subitem 2," & CStr(cnt)
     '   itmx.SubItems(3) = "subitem 3," & CStr(cnt)
     'Next

    .SortKey = 0
    .Sorted = False
    .View = lvwReport
    .GridLines = True
    .FullRowSelect = True
    .LabelEdit = lvwManual

  End With
  'Preethi - V1.2.2 20110301 -  IR8  - To solve the issue where the consultants can bypass the MI input
  'bolCheck = True
  'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
  'txtDataModified.Visible = False
  cmbContainer.Visible = False
'
Set gobjPNR = New CWT_GalileoPNR3.PNR
    gobjPNR.loadPNR


End Sub
Public Function incompleteMI() As String
    Dim intC As Integer
    Dim strMsg As String
        
    incompleteMI = ""
    With lvwClientMI
        If .ListItems.Count > 0 Then
            For intC = 1 To .ListItems.Count
                If .ListItems(intC).SubItems(2) = "" Then
                    strMsg = strMsg & "Incomplete MI for " & IIf(IsNumeric(.ListItems(intC).Text), "FF " & .ListItems(intC).Text, .ListItems(intC).Text) & "." & Chr(13)
                Else
                    If IsNumeric(.ListItems(intC).SubItems(4)) Then
                       If Len(.ListItems(intC).SubItems(2)) <> CInt(.ListItems(intC).SubItems(4)) Then
                          strMsg = strMsg & "Invalid data/Incompatible length detected for " & IIf(IsNumeric(.ListItems(intC).Text), "FF " & .ListItems(intC).Text, .ListItems(intC).Text) & "." & Chr(13)
                       Else
                          If UCase(.ListItems(intC).SubItems(3)) = "NUMERIC" And Not IsNumeric(fConvertZero(.ListItems(intC).SubItems(2))) Then
                             strMsg = strMsg & "Invalid data/Incompatible length detected for " & IIf(IsNumeric(.ListItems(intC).Text), "FF " & .ListItems(intC).Text, .ListItems(intC).Text) & "." & Chr(13)
                          End If
                       End If
                    Else
                        If UCase(.ListItems(intC).SubItems(3)) = "NUMERIC" And Not IsNumeric(fConvertZero(.ListItems(intC).SubItems(2))) Then
                            strMsg = strMsg & "Invalid data/Incompatible length detected for " & IIf(IsNumeric(.ListItems(intC).Text), "FF " & .ListItems(intC).Text, .ListItems(intC).Text) & "." & Chr(13)
                        End If
                    End If
                End If
            Next intC
        End If
    End With
    incompleteMI = strMsg
End Function
Private Function validMI(rowId As Integer) As Boolean
    Dim bolValid As Boolean
    
    bolValid = True
    With lvwClientMI
        If IsNumeric(.ListItems(rowId).SubItems(4)) Then
        'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
            'If Len(txtDataModified.Text) <> CInt(.ListItems(rowId).SubItems(4)) Then
            If Len(cmbClientMI.Text) <> CInt(.ListItems(rowId).SubItems(4)) Then
                bolValid = False
            End If
        End If
        'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
        'If UCase(.ListItems(rowId).SubItems(3)) = "NUMERIC" And Not IsNumeric(fConvertZero(txtDataModified.Text)) Then
        If UCase(.ListItems(rowId).SubItems(3)) = "NUMERIC" And Not IsNumeric(fConvertZero(cmbClientMI.Text)) Then
            bolValid = False
        End If
    End With
    validMI = bolValid
End Function
Public Sub pGetClientMI(CN As String)
    Dim rsMI As ADODB.Recordset
    Dim strSql As String
    Dim strSQL2 As String
    Dim item As ListItem
    'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
    Dim strFF As String
    Dim intC As Integer

    strSql = "select * from tblMICategory order by Code"
    Set rsMI = gdbConn.Execute(strSql)
    cmbMICat.AddItem "All"
    While Not rsMI.EOF
        If IsNumeric(rsMI!Code) Then
            cmbMICat.AddItem rsMI!Description, CInt(rsMI!Code)
        End If
        rsMI.MoveNext
    Wend
    If intLocation > 0 Then
        cmbMICat.listindex = intLocation
    Else
        cmbMICat.listindex = 0
    End If
    
    If UCase(cmbMICat.Text) = "ALL" Then
        strSql = "SELECT tblClientMI.*, tblMICategory.format FROM tblClientMI, tblMICategory where cn = '" & gobjPNR.CN & "' and tblClientMI.location = tblMICategory.code and isNumeric(FF)= 1 order by CAST(FF as integer)"
        strSQL2 = "SELECT tblClientMI.*, tblMICategory.format FROM tblClientMI, tblMICategory where cn = '" & gobjPNR.CN & "' and tblClientMI.location = tblMICategory.code and isNumeric(FF)= 0 order by FF"
    Else
        strSql = "SELECT tblClientMI.*, tblMICategory.format FROM tblClientMI, tblMICategory where cn = '" & gobjPNR.CN & "'" & _
                 " and location = " & cmbMICat.listindex & "  and tblClientMI.location = tblMICategory.code and isNumeric(FF)= 1 order by cast(FF as integer)"
        strSQL2 = "SELECT tblClientMI.*, tblMICategory.format FROM tblClientMI, tblMICategory where cn = '" & gobjPNR.CN & "'" & _
                 " and location = " & cmbMICat.listindex & "  and tblClientMI.location = tblMICategory.code and isNumeric(FF)= 0 order by FF"
    End If
    pRefreshData strSql, True
    pRefreshData strSQL2, False
    
    'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
    Set colFFValue = New Collection
    strFF = ""
    With lvwClientMI
        If .ListItems.Count > 0 Then
            For intC = 1 To .ListItems.Count
                If strFF = "" Then
                   strFF = "'" & .ListItems(intC).Text & "'"
                Else
                   strFF = strFF & ",'" & .ListItems(intC).Text & "'"
                End If
            Next
        End If
    End With
    
    Set colFFValue = GetClientMIValue(gobjPNR.CN, strFF)
    If colFFValue.Count = 0 Then
       cmbClientMI.style = fmStyleDropDownCombo
    Else
       cmbClientMI.style = fmStyleDropDownList
    End If
    
End Sub
Private Function pGetMIFormat(location As Integer) As String
    Dim rsMI As ADODB.Recordset
    Dim strSql As String
    'Retrieve formatting
    strSql = "select format from tblMICategory where code = " & CStr(location)
    Set rsMI = gdbConn.Execute(strSql)
    If Not rsMI.EOF Then
        pGetMIFormat = rsMI!Format
    Else
        pGetMIFormat = ""
    End If
    
End Function
Private Sub pRefreshData(sql As String, bolClear As Boolean)
    Dim rsMI As ADODB.Recordset
    Dim item As ListItem
    
    Set rsMI = gdbConn.Execute(sql)
    If bolClear = True Then
       lvwClientMI.ListItems.Clear
    End If
    While Not rsMI.EOF
        
        With lvwClientMI.ListItems
            Set item = .Add(, , IIf(IsNull(rsMI![FF]), "", rsMI![FF]))
            item.SubItems(1) = IIf(IsNull(rsMI![Description]), "", rsMI![Description])
            
             '230108
            If intLocation <> 6 And ((strPdtType <> "CT" And strPdtType <> "BT") And _
                 (intProdCode <> "35" And _
                 intProdCode <> "41" And _
                 intProdCode <> "50" And _
                 intProdCode <> "70")) Then

            If rsMI![Format] = "F" Then
                If intProdCode = "00" Then
                    item.SubItems(2) = pGetMIFromGDS(rsMI![FF], , "M", intProdCode)
                Else
                    item.SubItems(2) = pGetMIFromGDS(rsMI![FF], intFileFareNum, rsMI!Format)
                End If
            Else
                item.SubItems(2) = pGetMIFromGDS(rsMI![FF], , rsMI!Format, intProdCode)
 
            End If
            
            End If
            item.SubItems(3) = IIf(IsNull(rsMI![dataType]), "", rsMI![dataType])
            item.SubItems(4) = IIf(IsNull(rsMI![length]), "", rsMI![length])
        End With
        rsMI.MoveNext
    Wend
    
End Sub
Private Function getDIFreeFields() As String
    Dim intC As Integer
    Dim strDI As String
    Dim FFLineNo As Integer
    
    With lvwClientMI
    If .ListItems.Count > 0 Then
    For intC = 1 To .ListItems.Count
        If .ListItems(intC).SubItems(2) <> "" Then
            If .ListItems(intC).SubItems(2) <> pGetMIFromGDS(.ListItems(intC).Text, intFileFareNum, , , , FFLineNo) Then
                If .ListItems(intC).Text = "AC" Then
                    If FFLineNo > 0 Then
                        strDI = strDI & IIf(strDI <> "", "+", "") & "DI." & FFLineNo & "@AC-AAA." & IIf(intFileFareNum > 0 And IsNumeric(.ListItems(intC).Text), " * " & intFileFareNum & " / ", "") & .ListItems(intC).SubItems(2)
                    Else
                        strDI = strDI & IIf(strDI <> "", "+", "") & "DI.AC-AAA." & IIf(intFileFareNum > 0 And IsNumeric(.ListItems(intC).Text), " * " & intFileFareNum & " / ", "") & .ListItems(intC).SubItems(2)
                    End If
                Else
                    strDI = strDI & IIf(strDI <> "", "+", "") & IIf(IsNumeric(.ListItems(intC).Text), "DI.FT-FF", "DI.FT-") & .ListItems(intC).Text & "/" & IIf(intFileFareNum > 0 And IsNumeric(.ListItems(intC).Text), "*" & intFileFareNum & "/", "") & .ListItems(intC).SubItems(2)
                End If
            End If
        End If
    Next intC
    End If
    End With
                
    If strDI <> "" Then
        getDIFreeFields = strDI
    Else
        getDIFreeFields = ""
    End If
End Function
Public Function getMSXFreeFields() As String
    Dim intC As Integer
    Dim strDI As String
    With lvwClientMI
    If .ListItems.Count > 0 Then
    For intC = 1 To .ListItems.Count
        If .ListItems(intC).SubItems(2) <> "" Then
            'If .ListItems(intC).SubItems(2) <> pGetMIFromGDS(.ListItems(intC).Text, , "M", intProdCode) Then
                strDI = strDI & IIf(strDI <> "", "/", "") & .ListItems(intC).Text & "-" & .ListItems(intC).SubItems(2)
            'End If
        End If
    Next intC
    End If
    End With
    If strDI <> "" Then
        getMSXFreeFields = strDI
    Else
        getMSXFreeFields = ""
    End If

End Function
Private Sub Form_Unload(Cancel As Integer)

Call pResetValue

If pGetMIFormat(cmbMICat.listindex) = "D" Then
    Unload Me
    'pRedisplayMenu
Else
    Unload Me
End If

End Sub
Public Sub pResetValue()
   intLocation = Empty
   intFileFareNum = Empty
   intProdCode = Empty
   intPaxID = Empty
   MSXfreefields = Empty
   DIfreefields = Empty
End Sub
'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI

'Private Sub txtDataModified_KeyPress(KeyAscii As Integer)
'    With lvwClientMI
'    If .ListItems.Count > 0 Then
'        If IsNumeric(.SelectedItem.Index) Then
'            If .SelectedItem.Index > 0 Then
'                If UCase(.ListItems(.SelectedItem.Index).SubItems(3)) = "NUMERIC" Then
'                    KeyAscii = fAllowNumeric(KeyAscii, ".")
'                ElseIf UCase(.ListItems(.SelectedItem.Index).SubItems(3)) = "ALPHA" Then
'                    KeyAscii = fAllowAlpha(KeyAscii)
'                End If
'            End If
'        End If
'
'    End If
'    End With
'
'End Sub
'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI

'Private Sub txtDataModified_Validate(Cancel As Boolean)
'Dim strMsg As String
'If IsNumeric(lvwClientMI.SelectedItem.Index) Then
'    If Not validMI(lvwClientMI.SelectedItem.Index) Then
'        'MsgBox "Invalid data/Incompatible length detected.", vbCritical
'        strMsg = "Invalid data/Incompatible length detected."
'        modMsgBox.OKMsg = "OK"
'        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
'        txtDataModified.SetFocus
'    End If
'End If
'
'End Sub

Public Function fAllowAlpha(ByRef AsciiCode As Integer, Optional ByVal OtherAllowedCharacters As String = "") As Integer
Dim lngC As Long

    Select Case AsciiCode
           Case 8, 65 To 90, 97 To 122
               fAllowAlpha = Asc(UCase(Chr(AsciiCode))) ' change valid characters to uppercase
           Case Else
                For lngC = 1 To Len(OtherAllowedCharacters)
                    If Asc(Mid(OtherAllowedCharacters, lngC, 1)) = AsciiCode Then
                        fAllowAlpha = AsciiCode
                        Exit Function
                    End If
                Next
                fAllowAlpha = 0
    End Select

End Function
Public Function fAllowNumeric(ByRef AsciiCode As Integer, Optional ByVal OtherAllowedCharacters As String = "") As Integer
Dim lngC As Long

    Select Case AsciiCode
           Case 8, 48 To 57
               fAllowNumeric = Asc(UCase(Chr(AsciiCode))) ' change valid characters to uppercase USED IN CASE OTHERS ARE ALPHA
           Case Else
                For lngC = 1 To Len(OtherAllowedCharacters)
                    If Asc(Mid(OtherAllowedCharacters, lngC, 1)) = AsciiCode Then
                        fAllowNumeric = AsciiCode
                        Exit Function
                    End If
                Next

               fAllowNumeric = 0
    End Select

End Function

Private Sub lvwClientMI_ColumnClick(ByVal ColumnHeader As ColumnHeader)

'hide the text box
'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
  'txtDataModified.Visible = False
  cmbContainer.Visible = False

'sort the items
  lvwClientMI.SortKey = ColumnHeader.Index - 1
  lvwClientMI.SortOrder = Abs(Not lvwClientMI.SortOrder = 1)
  lvwClientMI.Sorted = True

End Sub
Private Sub lvwClientMI_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

'this routine:
'1. sets the last change if the dataModified flag is set
'2. sets a flag to prevent setting the dataModified flag
'3. determines the item or subitem clicked
'4. calc's the position for the text box
'5. moves and shows the text box
'6. clears the dataModified flag
'7. clears the DoingSetup flag

  Dim hti As LVHITTESTINFO
  Dim fpx As Single
  Dim fpy As Single
  Dim fpw As Single
  Dim fph As Single
  Dim rc As RECT
  Dim topindex As Long
  
  Dim intI As Integer

'prevent the textbox change event from
'registering as dataModified when the text is
'assigned to the textbox
  dataSetup = True

'if a pending dataModified flag is set, update the
'last edited item before moving on
  If dataModified And dwLastSubitemEdited > 0 Then
  'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
     'itmClicked.SubItems(dwLastSubitemEdited) = txtDataModified.Text
            itmClicked.SubItems(dwLastSubitemEdited) = cmbClientMI.Text
  End If

'hide the textbox
'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
  'txtDataModified.Visible = False
  cmbContainer.Visible = False
  
  'If cmbContainer.Visible = True Then cmbContainer.Visible = False

'get the position of the click
  With hti
     .pt.X = (X / Screen.TwipsPerPixelX)
     .pt.Y = (Y / Screen.TwipsPerPixelY)
     .flags = LVHT_ONITEM
  End With

'find out which subitem was clicked
  Call SendMessage(lvwClientMI.hwnd, _
                   LVM_SUBITEMHITTEST, _
                   0, hti)

'if on an item (HTI.lngItem <> -1) and
'the click occurred on the subitem
'column of interest (HTI.lngSubItem = 2 -
'which is column 3 (0-based)) move and
'show the textbox

If hti.lngSubItem = 2 Then

   If hti.lngItem <> -1 And hti.lngSubItem > 0 Then

    'prevent the listview label editing
    'from occurring if the control has
    'full row select set
     lvwClientMI.LabelEdit = lvwManual

    'determine the bounding rectangle
    'of the subitem column
     rc.Left = LVIR_LABEL
     rc.Top = hti.lngSubItem
     Call SendMessage(lvwClientMI.hwnd, _
                      LVM_GETSUBITEMRECT, _
                      hti.lngItem, _
                      rc)

    'we need to keep track of which
    'item was clicked so the item can
    'be updated later
    'position the text box
     Set itmClicked = lvwClientMI.ListItems(hti.lngItem + 1)
     itmClicked.Selected = True

    'get the current top index
     topindex = SendMessage(lvwClientMI.hwnd, _
                            LVM_GETTOPINDEX, _
                            0&, _
                            ByVal 0&)

    'establish the bounding rect for
    'the subitem in VB terms (the x
    'and y coordinates, and the height
    'and width of the item
     fpx = lvwClientMI.Left + _
             (rc.Left * Screen.TwipsPerPixelX) + 80

     fpy = lvwClientMI.Top + _
             (hti.lngItem + 1 - topindex) + _
             (rc.Top * Screen.TwipsPerPixelY)

    'a hard-coded height for the text box
     fph = 300

    'get the column width for the subitem
     fpw = SendMessage(lvwClientMI.hwnd, _
                       LVM_GETCOLUMNWIDTH, _
                       hti.lngSubItem, _
                       ByVal 0&)

    'calc the required width of
    'the textbox to fit in the column
     fpw = (fpw * Screen.TwipsPerPixelX) - 40

    'assign the current subitem
    'value to the textbox
      'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI

'     With txtDataModified
'
'        .Text = itmClicked.SubItems(hti.lngSubItem)
'
'        dwLastSubitemEdited = hti.lngSubItem
'
'       'position it over the subitem, make
'       'visible and assure the text box
'       'appears overtop the listview
'        .Move fpx, fpy, fpw, fph
'        .Visible = True
'        .ZOrder 0
'        .SetFocus
'
'     End With

With cmbContainer
   
   .Left = fpx
   .Top = fpy
   .Height = fph
    .Width = fpw
    .Visible = True
    .SetFocus
    .ZOrder 0
End With

PopulatecmbMI cmbClientMI, colFFValue, lvwClientMI.ListItems(hti.lngItem + 1).Text, itmClicked.SubItems(hti.lngSubItem)
If cmbClientMI.ListCount > 0 Then
   cmbClientMI.style = fmStyleDropDownList
   'cmbClientMI.listindex = 0
Else
   cmbClientMI.style = fmStyleDropDownCombo
End If

With cmbClientMI
   
   .Width = fpw
   .Height = fph
   AutoSizeDropDownWidth cmbClientMI
   If (Trim(itmClicked.SubItems(hti.lngSubItem)) <> "") Then
      .Text = itmClicked.SubItems(hti.lngSubItem)
   End If
   
   dwLastSubitemEdited = hti.lngSubItem
   
End With
    'clear the setup flag to allow the
    'textbox change event to set the
    '"dataModified" flag, and clear that flag
    'in preparation for editing
     dataSetup = False
     dataModified = False

  End If
End If
End Sub

Private Sub lvwClientMI_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

'if showing the text box, set
'focus to it and select any
'text in the control
'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
'  If txtDataModified.Visible = True Then
'
'     With txtDataModified
'        .SetFocus
'        .SelStart = 0
'        .SelLength = Len(.Text)
'     End With
'
'  End If
If cmbClientMI.Visible = True Then
   With cmbClientMI
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
   End With
End If

End Sub

'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
'Private Sub txtDataModified_Change()
'
'If Not dataSetup Then
'   dataModified = True
'End If
'
'
'End Sub

'Private Sub txtDataModified_LostFocus()
'
'If dataModified And dwLastSubitemEdited > 0 Then
'   itmClicked.SubItems(dwLastSubitemEdited) = txtDataModified.Text
'   dataModified = False
'End If


'End Sub

Private Sub updateDIField(ByRef strTemp As String)

Dim strDI() As String
Dim strFF() As String
Dim lngC As Integer
Dim strCmp1 As String
Dim strCmp2 As String
Dim i As Integer
Dim j As Integer

'Update the DI line that exists in PNR instead of adding new DI line
strDI = Split(strTemp, "+")
For lngC = LBound(strDI) To UBound(strDI)
    strFF = Split(strDI(lngC), "-")
    If UBound(strFF) > 0 Then
       With gobjPNR
            For j = 1 To .AcctRemarkCount
                i = InStr(1, .AcctRemark(j).RemarkText, "/")
                strCmp1 = Mid(.AcctRemark(j).RemarkText, 1, IIf(i > 0, i - 1, Len(.AcctRemark(j).RemarkText)))
                i = InStr(1, strFF(1), "/")
                strCmp2 = Mid(strFF(1), 1, IIf(i > 0, i - 1, Len(strFF(1))))
                If UCase(Trim(strCmp1)) = UCase(Trim(strCmp2)) Then
                   i = InStr(1, strDI(lngC), "DI.")
                   If i > 0 Then
                      strTemp = Replace(strTemp, strDI(lngC), "DI." & .AcctRemark(j).ItemNum & "@" & Mid(strDI(lngC), i + 3))
                      Exit For
                   End If
                End If
            Next j
       End With
    End If
Next
End Sub


