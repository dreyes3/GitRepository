VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeatMapRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarRowNumber As String 'local copy
Private mvarRowAttribute As SeatRowAttribute 'local copy
Private mvarSeat As SeatMapSeat 'local copy
Private mcolRowAttribute As Collection
Private mcolSeat As Collection



Public Property Let RowNumber(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RowNumber = 5
    mvarRowNumber = vData
End Property


Public Property Get RowNumber() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RowNumber
    RowNumber = mvarRowNumber
End Property



Public Sub RowAttributeAdd(ByVal RowAttribute As SeatRowAttribute)
    mcolRowAttribute.Add RowAttribute
    Set RowAttribute = Nothing
End Sub

Public Function RowAttribute(ByVal Index As Long) As SeatRowAttribute
    Set RowAttribute = mcolRowAttribute.Item(Index)
End Function

Public Property Get RowAttributeCount() As Long
    RowAttributeCount = mcolRowAttribute.Count
End Property


Public Sub SeatAdd(ByVal Seat As SeatMapSeat)
    mcolSeat.Add Seat
    Set Seat = Nothing
End Sub

Public Function Seat(ByVal Index As Long) As SeatMapSeat
    Set Seat = mcolSeat.Item(Index)
End Function

Public Property Get SeatCount() As Long
    SeatCount = mcolSeat.Count
End Property

Private Sub Class_Initialize()
Set mcolRowAttribute = New Collection
Set mcolSeat = New Collection
End Sub
