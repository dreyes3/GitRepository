VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrEmail As String 'local copy
Private mstrFaxNo As String 'local copy


Public Property Let EmailAddr(ByVal Email As String)
    mstrEmail = Email
End Property

Public Property Get EmailAddr() As String
    EmailAddr = mstrEmail
End Property

Public Property Let FaxNo(ByVal FaxNum As String)
    mstrFaxNo = FaxNum
End Property

Public Property Get FaxNo() As String
    FaxNo = mstrFaxNo
End Property




