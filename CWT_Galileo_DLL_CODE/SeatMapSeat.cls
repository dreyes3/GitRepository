VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeatMapSeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Private mvarSeatAttribute As SeatMapSeatAttribute 'local copy
Private mcolSeatAttribute As Collection


Private mvarColumnID As String 'local copy
Private mvarStatus As String 'local copy
Public Property Let Status(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Status = 5
    mvarStatus = vData
End Property


Public Property Get Status() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Status
    Status = mvarStatus
End Property



Public Property Let ColumnID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ColumnID = 5
    mvarColumnID = vData
End Property


Public Property Get ColumnID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ColumnID
    ColumnID = mvarColumnID
End Property


Public Sub SeatAttributeAdd(ByVal SeatAttribute As SeatMapSeatAttribute)
    mcolSeatAttribute.Add SeatAttribute
    Set SeatAttribute = Nothing
End Sub

Public Function SeatAttribute(ByVal Index As Long) As SeatMapSeatAttribute
    Set SeatAttribute = mcolSeatAttribute.Item(Index)
End Function

Public Property Get SeatAttributeCount() As Long
    SeatAttributeCount = mcolSeatAttribute.Count
End Property



Private Sub Class_Initialize()
Set mcolSeatAttribute = New Collection
End Sub
