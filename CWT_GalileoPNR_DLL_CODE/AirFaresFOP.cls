VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AirFaresFOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'local variable(s) to hold property value(s)
Private mvarBegLine As Long
Private mvarEndLine As Long
Private mvarFFNumber As String 'local copy
Private mvarFOPAmount As Single 'local copy
Private mvarFOPType As String 'local copy
Private mvarFOP_CCCode As String 'local copy
Private mvarFOP_CCNum As String 'local copy

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
Public Property Let FFNumber(ByVal vData As String)
    mvarFFNumber = vData
End Property

Public Property Get FFNumber() As String
    FFNumber = mvarFFNumber
End Property
Public Property Let FOPType(ByVal vData As String)
    mvarFOPType = vData
End Property

Public Property Get FOPType() As String
    FOPType = mvarFOPType
End Property
Public Property Let FOPAmount(ByVal vData As Single)
    mvarFOPAmount = vData
End Property

Public Property Get FOPAmount() As Single
    FOPAmount = mvarFOPAmount
End Property
Public Property Let FOP_CCCode(ByVal vData As String)
    mvarFOP_CCCode = vData
End Property

Public Property Get FOP_CCCode() As String
    FOP_CCCode = mvarFOP_CCCode
End Property
Public Property Let FOP_CCNum(ByVal vData As String)
    mvarFOP_CCNum = vData
End Property

Public Property Get FOP_CCNum() As String
    FOP_CCNum = mvarFOP_CCNum
End Property

