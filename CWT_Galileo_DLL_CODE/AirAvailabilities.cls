VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AirAvailabilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim mcolAA As Collection
Public Function Item(Index As Variant) As AirAvailability
Attribute Item.VB_UserMemId = 0
    Set Item = mcolAA.Item(Index)
    
End Function


Public Property Get Count() As Long
    Count = mcolAA.Count
End Property

Public Function AddAirAvailability(DepartDate As Date, Origin As String, _
    Destination As String, Optional NumberSeats As Byte = 1, _
    Optional GDSAvailType As String = "G") As AirAvailability

Dim objAA As AirAvailability
Set objAA = New AirAvailability

With objAA
    .AvailType = GDSAvailType
    .DeptDate = DepartDate
    .Dest = Destination
    .NumSeats = NumberSeats
    .Orig = Origin
End With

mcolAA.Add objAA
Set objAA = Nothing

End Function

Private Sub Class_Initialize()
Set mcolAA = New Collection
End Sub

Private Sub Class_Terminate()
Set mcolAA = Nothing
End Sub
