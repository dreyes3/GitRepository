VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GASalesRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarBegLine As Long
Private mvarEndLine As Long
Private mvarBaseAmount As Single 'local copy
Private mvarBaseCurrency As String 'local copy
Private mvarCommissionAmount As Single 'local copy
Private mvarCCNumber As String 'local copy
Private mvarCollectedAmount As Single 'local copy
Private mvarFOP As String 'local copy
Private mvarProductCode As String 'local copy
Private mvarPONumber As String 'local copy
Private mvarSellAmount As Single 'local copy
Private mvarGSTAmount As Single 'local copy
Private mvarTicketNumber As String 'local copy
Private mvarTax As Single 'local copy
Private mvarVendorCode As String 'local copy
Private mcolFreeFields As Collection
Private mvarDisplanyNo As Integer

'CC - V1.2.20 20130416 - CR212 - TF Per Passenger Logic
Private mvarPaxNum As String



Public Property Let DisplanyNo(ByVal vData As Integer)
    mvarDisplanyNo = vData
End Property

Public Property Get DisplanyNo() As Integer
    DisplanyNo = mvarDisplanyNo
End Property

Public Property Let EndLine(ByVal vData As Long)
    mvarEndLine = vData
End Property

Public Property Get EndLine() As Long
    EndLine = mvarEndLine
End Property

Public Property Let BegLine(ByVal vData As Long)
    mvarBegLine = vData
End Property

Public Property Get BegLine() As Long
    BegLine = mvarBegLine
End Property

Public Sub FreeFieldAdd(FFText As String, BegLine As Integer, EndLine As Integer)
Dim intI As Integer
Dim strFFNum As String
Dim strMsg As String


If InStr(FFText, "-") = 0 Then
    strFFNum = "FF0"
    
    For intI = 1 To mcolFreeFields.Count
     
        If InStr(mcolFreeFields.Item(intI), "-") = 0 Then
            'Set objPNR = objPNR
            'Call gPNR.ErrorAdd("Invalid FF format in MS line")
            'objPNR.ErrorAdd "Invalid FF format in MS line"
            'ReDim mcolError(5)
            'mcolError(1) = "Invalid FF format in MS line"
            strMsg = "Invalid FF format in MS line (missing '-')" & vbCrLf
            strMsg = strMsg & "DI Line: " & BegLine & " to " & EndLine & vbCrLf
            strMsg = strMsg & "Invalid text: " & mcolFreeFields.Item(intI)
            'MsgBox "Invalid FF format in MS line", , "Load PNR"
            'MsgBox strMsg, , "Load PNR"
            modMsgBox.OKMsg = "OK"
            modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
            Exit Sub
        End If
        
    Next
    
Else
    strFFNum = "FF" & Mid(FFText, 3, InStr(FFText, "-") - 1 - Len("FF"))
End If


    mcolFreeFields.Add FFText, strFFNum




End Sub

Public Function FreeFieldText(Index As Long) As String
FreeFieldText = mcolFreeFields.Item(Index)
End Function
Public Property Get FreeFieldTextCount() As Long
    FreeFieldTextCount = mcolFreeFields.Count
End Property
Public Property Let VendorCode(ByVal vData As String)
    mvarVendorCode = vData
End Property

Public Property Get VendorCode() As String
    VendorCode = mvarVendorCode
End Property

Public Property Let Tax(ByVal vData As Single)
    mvarTax = vData
End Property

Public Property Get Tax() As Single
    Tax = mvarTax
End Property

Public Property Let TicketNumber(ByVal vData As String)
    mvarTicketNumber = vData
End Property

Public Property Get TicketNumber() As String
    TicketNumber = mvarTicketNumber
End Property

Public Property Let SellAmount(ByVal vData As Single)
    mvarSellAmount = vData
End Property

Public Property Get SellAmount() As Single
    SellAmount = mvarSellAmount
End Property
Public Property Let GSTAmount(ByVal vData As Single)
    mvarGSTAmount = vData
End Property

Public Property Get GSTAmount() As Single
    GSTAmount = mvarGSTAmount
End Property

Public Property Let PONumber(ByVal vData As String)
    mvarPONumber = vData
End Property

Public Property Get PONumber() As String
    PONumber = mvarPONumber
End Property

Public Property Let ProductCode(ByVal vData As String)
    mvarProductCode = vData
End Property

Public Property Get ProductCode() As String
    ProductCode = mvarProductCode
End Property

Public Property Let FOP(ByVal vData As String)
    mvarFOP = vData
End Property

Public Property Get FOP() As String
    FOP = mvarFOP
End Property

Public Property Let CollectedAmount(ByVal vData As Single)
    mvarCollectedAmount = vData
End Property

Public Property Get CollectedAmount() As Single
    CollectedAmount = mvarCollectedAmount
End Property

Public Property Let CCNumber(ByVal vData As String)
    mvarCCNumber = vData
End Property

Public Property Get CCNumber() As String
    CCNumber = mvarCCNumber
End Property

Public Property Let CommissionAmount(ByVal vData As Single)
    mvarCommissionAmount = vData
End Property

Public Property Get CommissionAmount() As Single
    CommissionAmount = mvarCommissionAmount
End Property

Public Property Let BaseCurrency(ByVal vData As String)
    mvarBaseCurrency = vData
End Property

Public Property Get BaseCurrency() As String
    BaseCurrency = mvarBaseCurrency
End Property

Public Property Let BaseAmount(ByVal vData As Single)
    mvarBaseAmount = vData
End Property

Public Property Get BaseAmount() As Single
    BaseAmount = mvarBaseAmount
End Property

'CC - V1.2.20 20130416 - CR212 - TF Per Passenger Logic
Public Property Let PaxNum(ByVal vData As String)
    mvarPaxNum = vData
End Property

Public Property Get PaxNum() As String
    PaxNum = mvarPaxNum
End Property

Private Sub Class_Initialize()
Set mcolFreeFields = New Collection
End Sub

Private Sub Class_Terminate()
Set mcolFreeFields = Nothing
End Sub
