VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Remark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 'Preethi - V1.2.8  20111031 - CR101 - Aqua Itin Remark Screen
 
 Option Explicit

Private mvarNumber As Long
Private mvarRemarkType As String
Private mvarRemarkText As String
Private mvarDescription As String
Public Property Let Number(ByVal vData As Long)
    mvarNumber = vData
End Property

Public Property Get Number() As Long
    Number = mvarNumber
End Property

Public Property Let RemarkType(ByVal vData As String)
    mvarRemarkType = vData
End Property

Public Property Get RemarkType() As String
    RemarkType = mvarRemarkType
End Property
Public Property Let RemarkText(ByVal vData As String)
    mvarRemarkText = vData
End Property

Public Property Get RemarkText() As String
    RemarkText = mvarRemarkText
End Property
Public Property Let Description(ByVal vData As String)
    mvarDescription = vData
End Property

Public Property Get Description() As String
    Description = mvarDescription
End Property

