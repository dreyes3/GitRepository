VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmailAddresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolEmails As Collection

Public Sub AddEmail(ByVal Email As EmailAddress)
    mcolEmails.Add Email
End Sub

Public Property Get Email(Index As Integer) As EmailAddress
    Set Email = mcolEmails.item(Index)
End Property

Public Property Get EmailCount() As Integer
    EmailCount = mcolEmails.Count
End Property

Private Sub Class_Initialize()
    Set mcolEmails = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolEmails = Nothing
End Sub



